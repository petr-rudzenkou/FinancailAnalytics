using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Practices.Unity;
using FinancialAnalytics.Core.Composition.Unity;

namespace FinancialAnalytics.Core.Composition.Unity
{
    /// <summary>
    /// Instance of this class represents thread safe service locator with container.
    /// </summary>
    public class ServiceLocator :  IServiceLocator, IServiceContainer,Microsoft.Practices.ServiceLocation.IServiceLocator
    {
        private readonly IUnityContainer _container;

        //makes unity container thread safe for registering - temporal solution for registering in several places
        private readonly object _containerLock = new object();

        public ServiceLocator(IUnityContainer container)
        {
            if (container == null) throw new ArgumentNullException("container");   // invariant
            _container = container;
        }

        public static IServiceLocator Create()
        {
            var container = new ServiceContainer(new UnityContainer());
            return new ServiceLocator(container.GetInstance<IUnityContainer>());
        }

        public IServiceContainer Container
        {
            get
            {
                return this;
            }
        }

        public TService GetInstance<TService>(params ResolverOverride[] overrides)
        {
            return _container.Resolve<TService>(overrides);
        }

        public object GetService(Type serviceType)
        {
            return _container.Resolve(serviceType);
        }

        public T GetInstance<T>() 
        {
            return _container.Resolve<T>();
        }

        public object GetInstance(Type serviceType)
        {
            return _container.Resolve(serviceType);
        }


        public IEnumerable<T> GetInstances<T>()
        {
            return  _container.ResolveAll<T>();
        }

        public IEnumerable<TService> GetAllInstances<TService>()
        {
            return _container.ResolveAll<TService>();
        }

        public IEnumerable<object> GetAllInstances(Type serviceType)
        {
            return _container.ResolveAll(serviceType);
        }

        public TService GetInstance<TService>(string key)
        {
            return _container.Resolve<TService>(key);
        }

        public object GetInstance(Type serviceType, string key)
        {
            return _container.Resolve(serviceType, key);
        }

		public IEnumerable<object> GetInstances(Type serviceType)
		{
		    return _container.ResolveAll(serviceType);
		}

        IServiceContainer IServiceContainer.RegisterInstance<TInterface>(TInterface instance, LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TInterface>(_container,instance.GetType());
                _container.RegisterInstance<TInterface>(instance, lifetimeManager);
            }
            return this;
        }

        IServiceContainer IServiceContainer.Configure(params IConfigureContainer[] containerConfigurators)
        {
            foreach (var item in containerConfigurators)
            {
                item.ConfigureContainer(this);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>(LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>(lifetimeManager);
            }
            return this;
        }
        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>(string name)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>(name);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>(string name, LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>(name, lifetimeManager);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<TFrom, TTo>()
        {
            lock (_containerLock)
            {
                UnityContainerChecker.ThrowExceptionIfRegistered<TFrom>(_container, typeof(TTo));
                _container.RegisterType<TFrom, TTo>();
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterInstance<TInterface>(TInterface instance)
        {
            lock (_containerLock)
            {
                //ThrowExceptionIfRegistered<TInterface>(typeof(TInterface));
                _container.RegisterInstance<TInterface>(instance);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterInstance<TInterface>(string name, TInterface instance)
        {
            lock (_containerLock)
            {
                //ThrowExceptionIfRegistered<TInterface>(typeof(TInterface));
                _container.RegisterInstance<TInterface>(name, instance);
            }
            return this;
        }

        IServiceContainer IServiceContainer.RegisterType<T>(LifetimeManager lifetimeManager)
        {
            lock (_containerLock)
            {
                //ThrowExceptionIfRegistered<T>(typeof(T));
                _container.RegisterType<T>(lifetimeManager);
            }
            return this;
        }

        public IServiceContainer RegisterType<T>(LifetimeManager lifetimeManager, Func<IServiceLocator, T> factory)
        {
            lock (_containerLock)
            {
                //ThrowExceptionIfRegistered<T>(typeof(T));
                _container.RegisterType<T>(lifetimeManager,new InjectionFactory(x=> factory(this)));
            }
            return this;
        }

        public IServiceContainer RegisterType<T>(Func<IServiceLocator, T> factory)
        {
            lock (_containerLock)
            {
                //ThrowExceptionIfRegistered<T>(typeof(T));
                _container.RegisterType<T>(new InjectionFactory(x => factory(this)));
            }
            return this;
        }

        void IServiceContainer.Teardown(object instance)
        {
          _container.Teardown(instance);
        }

    }
}